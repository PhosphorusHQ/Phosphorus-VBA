Attribute VB_Name = "PPathTestsExpectedResults"
'@Folder PPathTests
Option Explicit

Public TestWorkbookName As String

Public Function TestPre() As String
  TestPre = ""
End Function

Public Function Test001() As String
  Test001 = ""
End Function

Public Function Test002() As String
  Test002 = ""
End Function

Public Function Test003() As String
  Test003 = ""
End Function

Public Function Test004() As String
  Test004 = ""
End Function

Public Function Test005() As String
  Test005 = _
    "/Pane[1][@Name='DropShadowTop']" & vbCrLf & _
    "/Pane[2][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']" & vbCrLf & _
    "/Pane[4][@Name='']" & vbCrLf & _
    "/Edit[1][@Name='Formula Bar']" & vbCrLf & _
    "/Pane[5][@Name='']" & vbCrLf & _
    "/TitleBar[1][@Name='']"
End Function

Public Function Test006() As String
  Test006 = ""
End Function

Public Function Test007() As String
  Test007 = ""
End Function

Public Function Test008() As String
  Test008 = Test005
End Function

Public Function Test009() As String

  Test009 = _
    "/Pane[1][@Name='DropShadowTop']" & vbCrLf & _
    "/Pane[2][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[1][@Name='Cell Mode Ready']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[2][@Name='Normal']"

  Test009 = Test009 & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[4][@Name='Page Break Preview']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[5][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[1][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[2][@Name='Zoom In']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[6][@Name='Zoom In']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[2][@Name='Zoom 10%']" & vbCrLf & _
    "/Pane[3][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']"
    
  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[1][@Name='AutoSave']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[2][@Name='Save']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/Button[1][@Name='Undo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/MenuItem[1][@Name='More Options']"
    
  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/Button[1][@Name='Redo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/MenuItem[1][@Name='Customize Quick Access Toolbar']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/TitleBar[1][@Name='Excel']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/TitleBar[1][@Name='Excel']/MenuItem[1][@Name='Excel']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']/Edit[1][@Name='Microsoft search']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[2][@Name='Lynn Gale']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[1][@Name='My Benefits']"
    
  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[2][@Name='Minimize']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[3][@Name='Restore Down']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[4][@Name='Close']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[5][@Name='File Tab']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[1][@Name='Home']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[2][@Name='Insert']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[4][@Name='Formulas']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[5][@Name='Data']"

  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[6][@Name='Review']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[7][@Name='View']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[8][@Name='Automate']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[9][@Name='Developer']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[10][@Name='Add-ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[11][@Name='Help']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[12][@Name='Acrobat']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[13][@Name='Power Pivot']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/Button[1][@Name='Comments']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/Button[2][@Name='Catch up']"
    
  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[6][@Name='Feedback']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/Button[1][@Name='Paste']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[1][@Name='Cut']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/Button[1][@Name='Copy']"
    
  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[2][@Name='Format Painter']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[3][@Name='Office Clipboard...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']/Edit[1][@Name='Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']/Edit[1][@Name='Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']/Button[1][@Name='Open']"
    
  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[1][@Name='Increase Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[2][@Name='Decrease Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[3][@Name='Bold']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[4][@Name='Italic']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/Button[1][@Name='Underline']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/Button[1][@Name='Bottom Border']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/MenuItem[1][@Name='More Options']"
    
  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/Button[1][@Name='Fill Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/Button[1][@Name='Font Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[5][@Name='Format Cell Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[1][@Name='Top Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[2][@Name='Middle Align']"
    
  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[3][@Name='Bottom Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[4][@Name='Align Left']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[5][@Name='Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[6][@Name='Align Right']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[7][@Name='Decrease Indent']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[8][@Name='Increase Indent']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[9][@Name='Wrap Text']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/Button[1][@Name='Merge & Center']"
    
  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[10][@Name='Format Cell Alignment']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']/Edit[1][@Name='Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/Button[1][@Name='Accounting Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[1][@Name='Percent Style']"
    
  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[2][@Name='Comma Style']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[3][@Name='Increase Decimal']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[4][@Name='Decrease Decimal']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[5][@Name='Format Cell Number']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[1][@Name='Conditional Formatting']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[2][@Name='Format as Table']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[3][@Name='Cell Styles']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']"

  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/Button[1][@Name='Insert Cells']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/Button[1][@Name='Delete Cells...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/MenuItem[1][@Name='Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/Button[1][@Name='Sum']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/MenuItem[1][@Name='More Options']"
  
  '141
  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[1][@Name='Fill']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[2][@Name='Clear']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[3][@Name='Sort & Filter']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[4][@Name='Find & Select']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[8][@Name='Add-ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[8][@Name='Add-ins']/Button[1][@Name='Add" & Excel.Application.WorksheetFunction.Unichar(8209) & "ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[9][@Name='Assistance']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[9][@Name='Assistance']/Button[1][@Name='Analyze Data']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[10][@Name='Adobe Acrobat']"
  
  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[10][@Name='Adobe Acrobat']/Button[1][@Name='Create a PDF']"

'Not always visible?
'  Test009_ExpectedAllElementsXpath = Test009_ExpectedAllElementsXpath & vbCrLf & _
'    "/Pane[@Name='']/ToolBar[@Name='']/Pane[@Name='Ribbon']/Pane[@Name='']/Pane[@Name='']/Pane[@Name='Ribbon']/Pane[@Name='Lower Ribbon']/Group[@Name='Home']/Group[@Name='Adobe Acrobat']/Button[@Name='Create a PDF and Share link']"

  Test009 = Test009 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[3][@Name='Ribbon Display Options']" & vbCrLf & _
    "/Pane[4][@Name='']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']/Edit[1][@Name='Name Box']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']/Button[1][@Name='Open']" & vbCrLf & _
    "/Edit[1][@Name='Formula Bar']" & vbCrLf & _
    "/Pane[5][@Name='']" & vbCrLf & _
    "/TitleBar[1][@Name='']"
    
  Test009 = Test009 & vbCrLf & _
    "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']" & vbCrLf & _
    "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']/MenuItem[1][@Name='System']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[1][@Name='Minimise']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[2][@Name='Restore']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[3][@Name='Close']"
    
'These have no runtimes - MSAA elements?
'  Test009_ExpectedAllElementsXpath = Test009_ExpectedAllElementsXpath & vbCrLf & _
    "/MenuBar[@Name='Worksheet Menu Bar']" & vbCrLf & _
    "/MenuBar[@Name='Worksheet Menu Bar']/ComboBox[@Name='Ask a Question']"

End Function

Public Function Test010() As String
  Test010 = Test009
End Function

Public Function Test011() As String
  Test011 = _
    "/Pane[1][@Name='DropShadowTop']" & vbCrLf & _
    "/Pane[2][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']" & vbCrLf & _
    "/Pane[4][@Name='']" & vbCrLf & _
    "/Pane[5][@Name='']"
End Function

Public Function Test012() As String
  Test012 = Test011
End Function

Public Function Test013() As String
  Test013 = _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/MenuItem[1][@Name='Customize Quick Access Toolbar']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/TitleBar[1][@Name='Excel']/MenuItem[1][@Name='Excel']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[2][@Name='Lynn Gale']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/MenuItem[1][@Name='More Options']"

  Test013 = Test013 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[1][@Name='Conditional Formatting']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[2][@Name='Format as Table']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[3][@Name='Cell Styles']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/MenuItem[1][@Name='Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/MenuItem[1][@Name='More Options']"
  
  Test013 = Test013 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[1][@Name='Fill']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[2][@Name='Clear']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[3][@Name='Sort & Filter']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[4][@Name='Find & Select']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[3][@Name='Ribbon Display Options']" & vbCrLf & _
    "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']/MenuItem[1][@Name='System']"

End Function

Public Function Test014() As String
  Test014 = Test005
End Function

Public Function Test015() As String
  Test015 = Test009
End Function

Public Function Test016() As String
  Test016 = "."
End Function

Public Function Test017() As String
  Test017 = Test016
End Function

Public Function Test018() As String
  Test018 = ""
End Function

Public Function Test019() As String
  Test019 = "/Edit[1][@Name='Formula Bar']/../Window[@Name='Excel']"
End Function

Public Function Test020() As String
  Test020 = Test019
End Function

Public Function Test021() As String
  Test021 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/../Pane[@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/../Pane[@Name='']/../Pane[@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/../Pane[@Name='']/../Pane[@Name='']/../Pane[@Name='Status Bar']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/../Pane[@Name='']/../Pane[@Name='']/../Pane[@Name='Status Bar']/../ToolBar[@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/../Pane[@Name='']/../Pane[@Name='']/../Pane[@Name='Status Bar']/../ToolBar[@Name='']/../Pane[@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/../Pane[@Name='']/../Pane[@Name='']/../Pane[@Name='Status Bar']/../ToolBar[@Name='']/../Pane[@Name='']/../Window[@Name='Excel']"
End Function

Public Function Test022() As String
  Test022 = _
    "/Edit[1][@Name='Formula Bar']/." & vbCrLf & _
    "/Edit[1][@Name='Formula Bar']/../Window[@Name='Excel']"
End Function

Public Function Test023() As String
  Test023 = _
    "." & vbCrLf & _
    Test015
End Function

Public Function Test024() As String
  Test024 = _
    "/Edit[1][@Name='Formula Bar']/../Pane[4][@Name='']" & vbCrLf & _
    "/Edit[1][@Name='Formula Bar']/../Pane[3][@Name='']" & vbCrLf & _
    "/Edit[1][@Name='Formula Bar']/../Pane[2][@Name='']" & vbCrLf & _
    "/Edit[1][@Name='Formula Bar']/../Pane[1][@Name='DropShadowTop']"
End Function

Public Function Test025() As String
  Test025 = _
    "/Edit[1][@Name='Formula Bar']/../Pane[5][@Name='']" & vbCrLf & _
    "/Edit[1][@Name='Formula Bar']/../TitleBar[1][@Name='']"
End Function

Public Function Test026() As String
  Test026 = ""
End Function

Public Function Test027() As String
    
'Not always visible?
'  Test027_ExpectedAllElementsXpath = Test027_ExpectedAllElementsXpath & vbCrLf & _
'    "/Pane[@Name='']/ToolBar[@Name='']/Pane[@Name='Ribbon']/Pane[@Name='']/Pane[@Name='']/Pane[@Name='Ribbon']/Pane[@Name='Lower Ribbon']/Group[@Name='Home']/Group[@Name='Adobe Acrobat']/Button[@Name='Create a PDF and Share link']"

  Test027 = _
    "/Pane[5][@Name='']" & vbCrLf & _
    "/Edit[1][@Name='Formula Bar']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']/Edit[1][@Name='Name Box']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']" & vbCrLf & _
    "/Pane[4][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[3][@Name='Ribbon Display Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[10][@Name='Adobe Acrobat']/Button[1][@Name='Create a PDF']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[10][@Name='Adobe Acrobat']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[9][@Name='Assistance']/Button[1][@Name='Analyze Data']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[9][@Name='Assistance']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[8][@Name='Add-ins']/Button[1][@Name='Add" & Excel.Application.WorksheetFunction.Unichar(8209) & "ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[8][@Name='Add-ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[4][@Name='Find & Select']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[3][@Name='Sort & Filter']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[2][@Name='Clear']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[1][@Name='Fill']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/Button[1][@Name='Sum']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/MenuItem[1][@Name='Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/MenuItem[1][@Name='More Options']"
    
  Test027 = Test027 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/Button[1][@Name='Delete Cells...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/Button[1][@Name='Insert Cells']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[3][@Name='Cell Styles']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[2][@Name='Format as Table']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[1][@Name='Conditional Formatting']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[5][@Name='Format Cell Number']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[4][@Name='Decrease Decimal']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[3][@Name='Increase Decimal']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[2][@Name='Comma Style']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[1][@Name='Percent Style']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/Button[1][@Name='Accounting Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']/Edit[1][@Name='Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[10][@Name='Format Cell Alignment']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/MenuItem[1][@Name='More Options']"

  Test027 = Test027 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/Button[1][@Name='Merge & Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[9][@Name='Wrap Text']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[8][@Name='Increase Indent']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[7][@Name='Decrease Indent']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[6][@Name='Align Right']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[5][@Name='Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[4][@Name='Align Left']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[3][@Name='Bottom Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[2][@Name='Middle Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[1][@Name='Top Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[5][@Name='Format Cell Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/Button[1][@Name='Font Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/Button[1][@Name='Fill Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/Button[1][@Name='Bottom Border']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/MenuItem[1][@Name='More Options']"

  Test027 = Test027 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/Button[1][@Name='Underline']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[4][@Name='Italic']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[3][@Name='Bold']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[2][@Name='Decrease Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[1][@Name='Increase Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']/Edit[1][@Name='Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']/Edit[1][@Name='Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[3][@Name='Office Clipboard...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[2][@Name='Format Painter']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/Button[1][@Name='Copy']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[1][@Name='Cut']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/Button[1][@Name='Paste']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']"

  Test027 = Test027 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[6][@Name='Feedback']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/Button[2][@Name='Catch up']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/Button[1][@Name='Comments']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[13][@Name='Power Pivot']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[12][@Name='Acrobat']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[11][@Name='Help']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[10][@Name='Add-ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[9][@Name='Developer']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[8][@Name='Automate']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[7][@Name='View']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[6][@Name='Review']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[5][@Name='Data']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[4][@Name='Formulas']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[2][@Name='Insert']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[1][@Name='Home']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[5][@Name='File Tab']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[4][@Name='Close']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[3][@Name='Restore Down']"
    
  Test027 = Test027 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[2][@Name='Minimize']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[1][@Name='My Benefits']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[2][@Name='Lynn Gale']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']/Edit[1][@Name='Microsoft search']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/TitleBar[1][@Name='Excel']/MenuItem[1][@Name='Excel']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/TitleBar[1][@Name='Excel']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/MenuItem[1][@Name='Customize Quick Access Toolbar']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/Button[1][@Name='Redo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/Button[1][@Name='Undo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[2][@Name='Save']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[1][@Name='AutoSave']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']"
    
  Test027 = Test027 & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[2][@Name='Zoom 10%']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[6][@Name='Zoom In']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[2][@Name='Zoom In']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[1][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[5][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[4][@Name='Page Break Preview']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[2][@Name='Normal']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[1][@Name='Cell Mode Ready']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']" & vbCrLf & _
    "/Pane[1][@Name='DropShadowTop']"

End Function

Public Function Test028() As String
  
  Test028 = _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/Button[1][@Name='Undo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[2][@Name='Save']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[1][@Name='AutoSave']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[2][@Name='Zoom 10%']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[6][@Name='Zoom In']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[2][@Name='Zoom In']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[1][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[5][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[4][@Name='Page Break Preview']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[2][@Name='Normal']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[1][@Name='Cell Mode Ready']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']"
    
  Test028 = Test028 & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']" & vbCrLf & _
    "/Pane[1][@Name='DropShadowTop']"

End Function

Public Function Test029() As String
  Test029 = ""
End Function

Public Function Test030() As String

  Test030 = _
    "/Pane[2][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[1][@Name='Cell Mode Ready']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[2][@Name='Normal']"

  Test030 = Test030 & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[4][@Name='Page Break Preview']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[5][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[1][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[2][@Name='Zoom In']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[6][@Name='Zoom In']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[2][@Name='Zoom 10%']" & vbCrLf & _
    "/Pane[3][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']"
    
  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[1][@Name='AutoSave']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[2][@Name='Save']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/Button[1][@Name='Undo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/MenuItem[1][@Name='More Options']"
    
  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/Button[1][@Name='Redo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/MenuItem[1][@Name='Customize Quick Access Toolbar']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/TitleBar[1][@Name='Excel']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/TitleBar[1][@Name='Excel']/MenuItem[1][@Name='Excel']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']/Edit[1][@Name='Microsoft search']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[2][@Name='Lynn Gale']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[1][@Name='My Benefits']"
    
  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[2][@Name='Minimize']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[3][@Name='Restore Down']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[4][@Name='Close']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[5][@Name='File Tab']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[1][@Name='Home']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[2][@Name='Insert']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[4][@Name='Formulas']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[5][@Name='Data']"

  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[6][@Name='Review']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[7][@Name='View']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[8][@Name='Automate']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[9][@Name='Developer']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[10][@Name='Add-ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[11][@Name='Help']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[12][@Name='Acrobat']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[13][@Name='Power Pivot']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/Button[1][@Name='Comments']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/Button[2][@Name='Catch up']"
    
  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[6][@Name='Feedback']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/Button[1][@Name='Paste']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[1][@Name='Cut']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/Button[1][@Name='Copy']"
    
  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[2][@Name='Format Painter']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[3][@Name='Office Clipboard...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']/Edit[1][@Name='Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']/Edit[1][@Name='Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']/Button[1][@Name='Open']"
    
  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[1][@Name='Increase Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[2][@Name='Decrease Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[3][@Name='Bold']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[4][@Name='Italic']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/Button[1][@Name='Underline']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/Button[1][@Name='Bottom Border']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/MenuItem[1][@Name='More Options']"
    
  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/Button[1][@Name='Fill Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/Button[1][@Name='Font Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[5][@Name='Format Cell Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[1][@Name='Top Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[2][@Name='Middle Align']"
    
  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[3][@Name='Bottom Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[4][@Name='Align Left']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[5][@Name='Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[6][@Name='Align Right']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[7][@Name='Decrease Indent']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[8][@Name='Increase Indent']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[9][@Name='Wrap Text']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/Button[1][@Name='Merge & Center']"
    
  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[10][@Name='Format Cell Alignment']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']/Edit[1][@Name='Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/Button[1][@Name='Accounting Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[1][@Name='Percent Style']"
    
  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[2][@Name='Comma Style']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[3][@Name='Increase Decimal']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[4][@Name='Decrease Decimal']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[5][@Name='Format Cell Number']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[1][@Name='Conditional Formatting']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[2][@Name='Format as Table']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[3][@Name='Cell Styles']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']"

  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/Button[1][@Name='Insert Cells']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/Button[1][@Name='Delete Cells...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/MenuItem[1][@Name='Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/Button[1][@Name='Sum']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/MenuItem[1][@Name='More Options']"
  
  '141
  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[1][@Name='Fill']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[2][@Name='Clear']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[3][@Name='Sort & Filter']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[4][@Name='Find & Select']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[8][@Name='Add-ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[8][@Name='Add-ins']/Button[1][@Name='Add" & Excel.Application.WorksheetFunction.Unichar(8209) & "ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[9][@Name='Assistance']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[9][@Name='Assistance']/Button[1][@Name='Analyze Data']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[10][@Name='Adobe Acrobat']"
  
  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[10][@Name='Adobe Acrobat']/Button[1][@Name='Create a PDF']"

'Not always visible?
'  Test030 = Test030 & vbCrLf & _
'    "/Pane[@Name='']/ToolBar[@Name='']/Pane[@Name='Ribbon']/Pane[@Name='']/Pane[@Name='']/Pane[@Name='Ribbon']/Pane[@Name='Lower Ribbon']/Group[@Name='Home']/Group[@Name='Adobe Acrobat']/Button[@Name='Create a PDF and Share link']"

  Test030 = Test030 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[3][@Name='Ribbon Display Options']" & vbCrLf & _
    "/Pane[4][@Name='']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']/Edit[1][@Name='Name Box']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']/Button[1][@Name='Open']" & vbCrLf & _
    "/Edit[1][@Name='Formula Bar']" & vbCrLf & _
    "/Pane[5][@Name='']" & vbCrLf & _
    "/TitleBar[1][@Name='']"
    
  Test030 = Test030 & vbCrLf & _
    "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']" & vbCrLf & _
    "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']/MenuItem[1][@Name='System']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[1][@Name='Minimise']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[2][@Name='Restore']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[3][@Name='Close']"
    
'These have no runtimes - MSAA elements?
'  Test030 = Test030 & vbCrLf & _
    "/MenuBar[@Name='Worksheet Menu Bar']" & vbCrLf & _
    "/MenuBar[@Name='Worksheet Menu Bar']/ComboBox[@Name='Ask a Question']"

End Function

Public Function Test031() As String
      
  Test031 = _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/Button[1][@Name='Redo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/MenuItem[1][@Name='Customize Quick Access Toolbar']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/TitleBar[1][@Name='Excel']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/TitleBar[1][@Name='Excel']/MenuItem[1][@Name='Excel']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']/Edit[1][@Name='Microsoft search']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[2][@Name='Lynn Gale']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[1][@Name='My Benefits']"
    
  Test031 = Test031 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[2][@Name='Minimize']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[3][@Name='Restore Down']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[4][@Name='Close']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[5][@Name='File Tab']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[1][@Name='Home']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[2][@Name='Insert']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[4][@Name='Formulas']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[5][@Name='Data']"

  Test031 = Test031 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[6][@Name='Review']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[7][@Name='View']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[8][@Name='Automate']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[9][@Name='Developer']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[10][@Name='Add-ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[11][@Name='Help']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[12][@Name='Acrobat']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[13][@Name='Power Pivot']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/Button[1][@Name='Comments']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/Button[2][@Name='Catch up']"
    
  Test031 = Test031 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[6][@Name='Feedback']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/Button[1][@Name='Paste']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[1][@Name='Cut']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/Button[1][@Name='Copy']"
    
  Test031 = Test031 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[2][@Name='Format Painter']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[3][@Name='Office Clipboard...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']/Edit[1][@Name='Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']/Edit[1][@Name='Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']/Button[1][@Name='Open']"
    
  Test031 = Test031 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[1][@Name='Increase Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[2][@Name='Decrease Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[3][@Name='Bold']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[4][@Name='Italic']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/Button[1][@Name='Underline']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/Button[1][@Name='Bottom Border']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/MenuItem[1][@Name='More Options']"
    
  Test031 = Test031 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/Button[1][@Name='Fill Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/Button[1][@Name='Font Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[5][@Name='Format Cell Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[1][@Name='Top Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[2][@Name='Middle Align']"
    
  Test031 = Test031 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[3][@Name='Bottom Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[4][@Name='Align Left']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[5][@Name='Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[6][@Name='Align Right']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[7][@Name='Decrease Indent']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[8][@Name='Increase Indent']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[9][@Name='Wrap Text']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/Button[1][@Name='Merge & Center']"
    
  Test031 = Test031 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[10][@Name='Format Cell Alignment']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']/Edit[1][@Name='Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/Button[1][@Name='Accounting Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[1][@Name='Percent Style']"
    
  Test031 = Test031 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[2][@Name='Comma Style']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[3][@Name='Increase Decimal']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[4][@Name='Decrease Decimal']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[5][@Name='Format Cell Number']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[1][@Name='Conditional Formatting']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[2][@Name='Format as Table']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[3][@Name='Cell Styles']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']"

  Test031 = Test031 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/Button[1][@Name='Insert Cells']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/Button[1][@Name='Delete Cells...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/MenuItem[1][@Name='Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/Button[1][@Name='Sum']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/MenuItem[1][@Name='More Options']"
  
  '141
  Test031 = Test031 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[1][@Name='Fill']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[2][@Name='Clear']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[3][@Name='Sort & Filter']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[4][@Name='Find & Select']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[8][@Name='Add-ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[8][@Name='Add-ins']/Button[1][@Name='Add" & Excel.Application.WorksheetFunction.Unichar(8209) & "ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[9][@Name='Assistance']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[9][@Name='Assistance']/Button[1][@Name='Analyze Data']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[10][@Name='Adobe Acrobat']"
  
  Test031 = Test031 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[10][@Name='Adobe Acrobat']/Button[1][@Name='Create a PDF']"

'Not always visible?
'  Test031 = Test031 & vbCrLf & _
'    "/Pane[@Name='']/ToolBar[@Name='']/Pane[@Name='Ribbon']/Pane[@Name='']/Pane[@Name='']/Pane[@Name='Ribbon']/Pane[@Name='Lower Ribbon']/Group[@Name='Home']/Group[@Name='Adobe Acrobat']/Button[@Name='Create a PDF and Share link']"

  Test031 = Test031 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[3][@Name='Ribbon Display Options']" & vbCrLf & _
    "/Pane[4][@Name='']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']/Edit[1][@Name='Name Box']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']/Button[1][@Name='Open']" & vbCrLf & _
    "/Edit[1][@Name='Formula Bar']" & vbCrLf & _
    "/Pane[5][@Name='']" & vbCrLf & _
    "/TitleBar[1][@Name='']"
    
  Test031 = Test031 & vbCrLf & _
    "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']" & vbCrLf & _
    "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']/MenuItem[1][@Name='System']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[1][@Name='Minimise']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[2][@Name='Restore']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[3][@Name='Close']"
    
'These have no runtimes - MSAA elements?
'  Test031 = Test031 & vbCrLf & _
    "/MenuBar[@Name='Worksheet Menu Bar']" & vbCrLf & _
    "/MenuBar[@Name='Worksheet Menu Bar']/ComboBox[@Name='Ask a Question']"

End Function

Public Function Test032() As String
  Test032 = _
    "/Pane[1][@Name='DropShadowTop']" & vbCrLf & _
    "/Pane[2][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']" & vbCrLf & _
    "/Pane[4][@Name='']" & vbCrLf & _
    "/Pane[5][@Name='']"
End Function

Public Function Test033() As String

  '1
  Test033 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[2][@Name='Normal']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[4][@Name='Page Break Preview']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[5][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[1][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[2][@Name='Zoom In']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[6][@Name='Zoom In']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[1][@Name='AutoSave']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[2][@Name='Save']"

  '11
  Test033 = Test033 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/Button[1][@Name='Undo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/Button[1][@Name='Redo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[1][@Name='My Benefits']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[2][@Name='Minimize']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[3][@Name='Restore Down']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[4][@Name='Close']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[5][@Name='File Tab']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/Button[1][@Name='Comments']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/Button[2][@Name='Catch up']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[6][@Name='Feedback']"

  '21
  Test033 = Test033 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/Button[1][@Name='Paste']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[1][@Name='Cut']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/Button[1][@Name='Copy']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[2][@Name='Format Painter']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[3][@Name='Office Clipboard...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[1][@Name='Increase Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[2][@Name='Decrease Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[3][@Name='Bold']"

  '31
  Test033 = Test033 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[4][@Name='Italic']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/Button[1][@Name='Underline']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/Button[1][@Name='Bottom Border']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/Button[1][@Name='Fill Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/Button[1][@Name='Font Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[5][@Name='Format Cell Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[1][@Name='Top Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[2][@Name='Middle Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[3][@Name='Bottom Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[4][@Name='Align Left']"

  '41
  Test033 = Test033 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[5][@Name='Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[6][@Name='Align Right']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[7][@Name='Decrease Indent']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[8][@Name='Increase Indent']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[9][@Name='Wrap Text']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/Button[1][@Name='Merge & Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[10][@Name='Format Cell Alignment']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/Button[1][@Name='Accounting Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[1][@Name='Percent Style']"
    
  '51
  Test033 = Test033 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[2][@Name='Comma Style']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[3][@Name='Increase Decimal']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[4][@Name='Decrease Decimal']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[5][@Name='Format Cell Number']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/Button[1][@Name='Insert Cells']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/Button[1][@Name='Delete Cells...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/Button[1][@Name='Sum']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[8][@Name='Add-ins']/Button[1][@Name='Add" & Excel.Application.WorksheetFunction.Unichar(8209) & "ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[9][@Name='Assistance']/Button[1][@Name='Analyze Data']"

  '61
  Test033 = Test033 & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[10][@Name='Adobe Acrobat']/Button[1][@Name='Create a PDF']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']/Button[1][@Name='Open']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[1][@Name='Minimise']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[2][@Name='Restore']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[3][@Name='Close']"
    
End Function

Public Function Test034() As String
  Test034 = Test016
End Function

Public Function Test035() As String
  Test035 = Test016
End Function

Public Function Test036() As String
  Test036 = ""
End Function

Public Function Test037() As String
  Test037 = "/Edit[1][@Name='Formula Bar']/../Window[@Name='Excel']"
End Function

Public Function Test038() As String
  Test038 = "/Edit[1][@Name='Formula Bar']/../Window[@Name='Excel']"
End Function

Public Function Test039() As String
'TODO: Check this result!
  Test039 = "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/../Pane[@Name='']/../Pane[@Name='']/../Pane[@Name='Status Bar']/../ToolBar[@Name='']"
End Function

Public Function Test040() As String
  Test040 = "/Edit[1][@Name='Formula Bar']/../Window[@Name='Excel']"
End Function

Public Function Test041() As String
  Test041 = _
    "/Pane[3][@Name='']/." & vbCrLf & _
    "/Pane[3][@Name='']//ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']//ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']//ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']//ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']" & vbCrLf & _
    "/Pane[3][@Name='']//ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']"

End Function

Public Function Test042() As String
  Test042 = _
  "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/../Button[2][@Name='Save']" & vbCrLf & _
  "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/../Button[1][@Name='AutoSave']"
End Function

Public Function Test043() As String
  Test043 = "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/../Slider[1][@Name='Zoom']"
End Function

Public Function Test044() As String
  Test044 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[5][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[4][@Name='Page Break Preview']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[2][@Name='Normal']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']"
End Function

Public Function Test045() As String
  Test045 = _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test046() As String
  
  Test046 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@BoundingRectangle='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@ClassName='NetUISimpleButton'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@Culture=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@FillType=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@FrameworkId='Win32'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@FullDescription='No macros are currently recording. Click to begin recording a new macro.'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@HasKeyboardFocus=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@HeadingLevel=80050" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@IsContentElement=True" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@IsControlElement=True" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@IsDataValidForForm=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@IsDialog=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@IsEnabled=True" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@IsKeyboardFocusable=True" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@IsOffscreen=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@IsPassword=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@IsPeripheral=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@IsRequiredForForm=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@LandmarkType=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@Level=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@LiveSetting=2" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@LocalizedControlType='Button'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@Name='Macro Recording Not Recording'" & vbCrLf
    
  Test046 = Test046 & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@NativeWindowHandle='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@OptimizeForVisualContent=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@Orientation=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@PositionInSet=2" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@ProcessId='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@ProviderDescription='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@SizeOfSet=2" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@VisualEffects=0"
    
End Function

Public Function Test047() As String
  Test047 = "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']/@IsEnabled=True"
End Function

Public Function Test048() As String
  
  Test048 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@BoundingRectangle='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@ClassName='NetUIRepeatButton'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@Culture=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@FillType=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@FrameworkId='Win32'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@FullDescription='Moves the horizontal position left a couple of columns'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@HasKeyboardFocus=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@HeadingLevel=80050" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@IsContentElement=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@IsControlElement=True" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@IsDataValidForForm=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@IsDialog=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@IsEnabled=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@IsKeyboardFocusable=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@IsOffscreen=True" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@IsPassword=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@IsPeripheral=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@IsRequiredForForm=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@LandmarkType=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@Level=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@LiveSetting=2" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@LocalizedControlType='Button'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@Name='Zoom Out'" & vbCrLf
    
  Test048 = Test048 & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@NativeWindowHandle='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@OptimizeForVisualContent=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@Orientation=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@PositionInSet=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@ProcessId='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@ProviderDescription='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@SizeOfSet=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/@VisualEffects=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@BoundingRectangle='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@ClassName='NetUIRepeatButton'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@Culture=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@FillType=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@FrameworkId='Win32'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@FullDescription='Moves the horizontal position right a couple of columns'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@HasKeyboardFocus=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@HeadingLevel=80050" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@IsContentElement=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@IsControlElement=True" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@IsDataValidForForm=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@IsDialog=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@IsEnabled=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@IsKeyboardFocusable=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@IsOffscreen=False" & vbCrLf
    
  Test048 = Test048 & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@IsPassword=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@IsPeripheral=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@IsRequiredForForm=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@LandmarkType=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@Level=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@LiveSetting=2" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@LocalizedControlType='Button'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@Name='Zoom In'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@NativeWindowHandle='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@OptimizeForVisualContent=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@Orientation=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@PositionInSet=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@ProcessId='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@ProviderDescription='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@SizeOfSet=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/@VisualEffects=0"

End Function

Public Function Test049() As String
  Test049 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']"
End Function

Public Function Test050() As String
  Test050 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[1][@Name='Zoom Out']/text()" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']//Button[2][@Name='Zoom In']/text()"
End Function

Public Function Test051() As String
  Test051 = ""
End Function

Public Function Test052() As String
  Test052 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[5][@Name='Zoom Out']/@value" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/@value" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[6][@Name='Zoom In']/@value"
End Function

Public Function Test053() As String
  Test053 = _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']/Edit[1][@Name='Microsoft search']/@value" & vbCrLf & _
    "/TitleBar[1][@Name='']/@value"
End Function

Public Function Test054() As String
  Test054 = _
    "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[1][@Name='Minimise']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[2][@Name='Restore']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[3][@Name='Close']"
End Function

Public Function Test055() As String
  Test055 = _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/Button[1][@Name='Undo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/Button[1][@Name='Redo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/Button[1][@Name='Paste']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/Button[1][@Name='Copy']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/Button[1][@Name='Underline']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/Button[1][@Name='Bottom Border']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/Button[1][@Name='Fill Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/Button[1][@Name='Font Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/Button[1][@Name='Merge & Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/Button[1][@Name='Accounting Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/Button[1][@Name='Insert Cells']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/Button[1][@Name='Delete Cells...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/Button[1][@Name='Sum']"
End Function

Public Function Test056() As String
  Test056 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@BoundingRectangle='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@ClassName='MsoCommandBar'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@Culture=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@FillType=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@FrameworkId='Win32'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@HasKeyboardFocus=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@HeadingLevel=80050" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@HelpText='Status Bar toolbar'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@IsContentElement=True" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@IsControlElement=True" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@IsDataValidForForm=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@IsDialog=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@IsEnabled=True" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@IsKeyboardFocusable=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@IsOffscreen=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@IsPassword=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@IsPeripheral=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@IsRequiredForForm=False" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@LandmarkType=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@Level=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@LiveSetting=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@LocalizedControlType='tool bar'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@NativeWindowHandle='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@OptimizeForVisualContent=False" & vbCrLf
    
  Test056 = Test056 & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@Orientation=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@PositionInSet=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@ProcessId='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@ProviderDescription='#'" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@SizeOfSet=0" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/@VisualEffects=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@BoundingRectangle='#'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@ClassName='MsoCommandBar'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@Culture=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@FillType=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@FrameworkId='Win32'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@HasKeyboardFocus=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@HeadingLevel=80050" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@HelpText='Ribbon toolbar'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@IsContentElement=True" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@IsControlElement=True" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@IsDataValidForForm=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@IsDialog=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@IsEnabled=True" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@IsKeyboardFocusable=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@IsOffscreen=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@IsPassword=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@IsPeripheral=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@IsRequiredForForm=False" & vbCrLf
    
  Test056 = Test056 & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@LandmarkType=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@Level=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@LiveSetting=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@LocalizedControlType='tool bar'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@NativeWindowHandle='#'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@OptimizeForVisualContent=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@Orientation=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@PositionInSet=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@ProcessId='#'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@ProviderDescription='#'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@SizeOfSet=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/@VisualEffects=0" & vbCrLf

  Test056 = Test056 & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@BoundingRectangle='#'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@ClassName='NetUIElement'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@Culture=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@FillType=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@FrameworkId='Win32'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@HasKeyboardFocus=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@HeadingLevel=80050" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@IsContentElement=True" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@IsControlElement=True" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@IsDataValidForForm=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@IsDialog=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@IsEnabled=True" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@IsKeyboardFocusable=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@IsOffscreen=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@IsPassword=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@IsPeripheral=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@IsRequiredForForm=False" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@LandmarkType=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@Level=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@LiveSetting=2" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@LocalizedControlType='ToolBar'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@Name='Quick Access Toolbar'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@NativeWindowHandle='#'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@OptimizeForVisualContent=False" & vbCrLf
    
  Test056 = Test056 & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@Orientation=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@PositionInSet=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@ProcessId='#'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@ProviderDescription='#'" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@SizeOfSet=0" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/@VisualEffects=0"
    
End Function

Public Function Test057() As String
  Test057 = "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']/@AutomationId='MenuBar'"
End Function

Public Function Test058() As String
  Test058 = _
    "/Pane[1][@Name='DropShadowTop']" & vbCrLf & _
    "/Pane[2][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']" & vbCrLf & _
    "/Pane[4][@Name='']" & vbCrLf & _
    "/Pane[5][@Name='']" & vbCrLf & _
    "/Edit[1][@Name='Formula Bar']"
End Function

Public Function Test059() As String
  Test059 = _
    "/Pane[1][@Name='DropShadowTop']" & vbCrLf & _
    "/Pane[2][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']" & vbCrLf & _
    "/Pane[4][@Name='']" & vbCrLf & _
    "/Pane[5][@Name='']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[1][@Name='Home']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[2][@Name='Insert']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[4][@Name='Formulas']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[5][@Name='Data']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[6][@Name='Review']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[7][@Name='View']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[8][@Name='Automate']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[9][@Name='Developer']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[10][@Name='Add-ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[11][@Name='Help']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[12][@Name='Acrobat']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[13][@Name='Power Pivot']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//Button[1][@Name='Comments']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//Button[2][@Name='Catch up']"
End Function

Public Function Test060() As String
  
  Test060 = _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/MenuItem[1][@Name='Customize Quick Access Toolbar']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/TitleBar[1][@Name='Excel']/MenuItem[1][@Name='Excel']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[2][@Name='Lynn Gale']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[1][@Name='Conditional Formatting']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[2][@Name='Format as Table']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[3][@Name='Cell Styles']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/MenuItem[1][@Name='More Options']" & vbCrLf
  
  Test060 = Test060 & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/MenuItem[1][@Name='Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[1][@Name='Fill']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[2][@Name='Clear']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[3][@Name='Sort & Filter']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[4][@Name='Find & Select']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[3][@Name='Ribbon Display Options']"


End Function

Public Function Test061() As String
  Test061 = ""
End Function

Public Function Test062() As String
  Test062 = _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[1][@Name='Home']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[2][@Name='Insert']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[4][@Name='Formulas']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[5][@Name='Data']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[6][@Name='Review']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[7][@Name='View']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[8][@Name='Automate']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[9][@Name='Developer']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[10][@Name='Add-ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[11][@Name='Help']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[12][@Name='Acrobat']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//TabItem[13][@Name='Power Pivot']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//Button[1][@Name='Comments']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']//Button[2][@Name='Catch up']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[1][@Name='Cell Mode Ready']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[2][@Name='Normal']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[4][@Name='Page Break Preview']" & vbCrLf
   
  Test062 = Test062 & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[5][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[1][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[2][@Name='Zoom In']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[6][@Name='Zoom In']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']//Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[2][@Name='Zoom 10%']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']//Button[1][@Name='Undo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']//MenuItem[1][@Name='More Options']"

End Function

Public Function Test063() As String
  Test063 = "/Pane[2][@Name='']"
End Function

Public Function Test064() As String
  Test064 = "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']"
End Function

Public Function Test065() As String
  Test065 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[2][@Name='Normal']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[2][@Name='Zoom In']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[2][@Name='Save']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[2][@Name='Minimize']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/Button[2][@Name='Catch up']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[3][@Name='Office Clipboard...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[2][@Name='Decrease Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[2][@Name='Middle Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[2][@Name='Comma Style']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[2][@Name='Restore']"
End Function

Public Function Test066() As String
  Test066 = "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[2][@Name='Minimize']"
End Function

Public Function Test067() As String
  Test067 = "/Pane[2][@Name='']"
End Function

Public Function Test068() As String
  Test068 = "/Pane[1][@Name='DropShadowTop']"
End Function

Public Function Test069() As String
  Test069 = "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[2][@Name='Normal']"
End Function

Public Function Test070() As String
  Test070 = "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[2][@Name='Zoom In']"
End Function

Public Function Test071() As String
  Test071 = "/Pane[5][@Name='']"
End Function

Public Function Test072() As String
  Test072 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[5][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[2][@Name='Zoom In']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[6][@Name='Zoom In']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[2][@Name='Save']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/Button[1][@Name='Undo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/Button[1][@Name='Redo']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[5][@Name='File Tab']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/Button[2][@Name='Catch up']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[6][@Name='Feedback']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/Button[1][@Name='Paste']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[1][@Name='Cut']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/Button[1][@Name='Copy']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[3][@Name='Office Clipboard...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[4][@Name='Italic']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/Button[1][@Name='Underline']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/Button[1][@Name='Bottom Border']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/Button[1][@Name='Fill Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/Button[1][@Name='Font Color']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[5][@Name='Format Cell Font']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[9][@Name='Wrap Text']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/Button[1][@Name='Merge & Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[10][@Name='Format Cell Alignment']" & vbCrLf
    
  Test072 = Test072 & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']/Button[1][@Name='Open']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/Button[1][@Name='Accounting Number Format']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[5][@Name='Format Cell Number']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/Button[1][@Name='Insert Cells']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/Button[1][@Name='Delete Cells...']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/Button[1][@Name='Sum']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[8][@Name='Add-ins']/Button[1][@Name='Add" & Excel.Application.WorksheetFunction.Unichar(8209) & "ins']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[9][@Name='Assistance']/Button[1][@Name='Analyze Data']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[10][@Name='Adobe Acrobat']/Button[1][@Name='Create a PDF']" & vbCrLf & _
    "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']/Button[1][@Name='Open']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[3][@Name='Close']"
    
End Function

Public Function Test073() As String
  Test073 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[4][@Name='Page Break Preview']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[1][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[1][@Name='AutoSave']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[4][@Name='Close']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/Button[1][@Name='Comments']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[2][@Name='Format Painter']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[3][@Name='Bold']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[8][@Name='Increase Indent']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[4][@Name='Decrease Decimal']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[2][@Name='Restore']"
End Function

Public Function Test074() As String
  Test074 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[2][@Name='Normal']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[4][@Name='Page Break Preview']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[2][@Name='Minimize']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[3][@Name='Restore Down']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Button[4][@Name='Close']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[2][@Name='Decrease Font Size']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[3][@Name='Bold']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[2][@Name='Middle Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[3][@Name='Bottom Align']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[4][@Name='Align Left']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[5][@Name='Center']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[6][@Name='Align Right']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[7][@Name='Decrease Indent']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[8][@Name='Increase Indent']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[2][@Name='Comma Style']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[3][@Name='Increase Decimal']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/Button[4][@Name='Decrease Decimal']" & vbCrLf & _
    "/TitleBar[1][@Name='']/Button[2][@Name='Restore']"
End Function

Public Function Test075() As String
  Test075 = "/TitleBar[1][@Name='']/Button[1][@Name='Minimise']"
End Function

Public Function Test076() As String
  Test076 = _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[2][@Name='Normal']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[3][@Name='Page Layout']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[4][@Name='Page Break Preview']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[5][@Name='Zoom Out']" & vbCrLf & _
    "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[6][@Name='Zoom In']"
End Function

Public Function Test077() As String
  Test077 = "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']"
End Function

Public Function Test078() As String
End Function

Public Function Test079() As String
End Function

Public Function Test080() As String
  Test080 = "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']"
End Function

Public Function Test081() As String
  Test081 = "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']"
End Function

Public Function Test082() As String
  Test082 = "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']/Edit[1][@Name='Name Box']"
End Function

Public Function Test083() As String
  Test083 = "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']/Edit[1][@Name='Name Box']"
End Function

Public Function Test084() As String
  Test084 = _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/MenuItem[1][@Name='Customize Quick Access Toolbar']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/MenuItem[1][@Name='More Options']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[3][@Name='Ribbon Display Options']"
  
End Function

Public Function Test085() As String
  Test085 = _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[2][@Name='Lynn Gale']" & vbCrLf & _
    "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']"
End Function

Public Function Test086() As String
  Test086 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/MenuItem[1][@Name='Customize Quick Access Toolbar']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[2][@Name='Lynn Gale']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[3][@Name='Ribbon Display Options']"
End Function

Public Function Test087() As String
  Test087 = "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']"
End Function

Public Function Test088() As String
  Test088 = "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[2][@Name='Save']"
End Function

Public Function Test089() As String
  Test089 = "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']/Button[1][@Name='Open']"
End Function

Public Function Test090() As String
  Test090 = _
   "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[1][@Name='Cell Mode Ready']" & vbCrLf & _
   "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[2][@Name='Zoom 10%']"
End Function

Public Function Test091() As String
  Test091 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
   "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test092() As String
  Test092 = "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[1][@Name='AutoSave']"
End Function

Public Function Test093() As String
  Test093 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[1][@Name='Home']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[2][@Name='Insert']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[3][@Name='Page Layout']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[4][@Name='Formulas']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[5][@Name='Data']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[6][@Name='Review']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[7][@Name='View']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[8][@Name='Automate']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[9][@Name='Developer']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[10][@Name='Add-ins']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[11][@Name='Help']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[12][@Name='Acrobat']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[13][@Name='Power Pivot']"
End Function

Public Function Test094() As String
  Test094 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
   "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test095() As String
  Test095 = "/Pane[1][@Name='DropShadowTop']"
End Function

Public Function Test096() As String
  Test096 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
   "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test097() As String
  Test097 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
   "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test098() As String
  Test098 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
   "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test099() As String
  Test099 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
   "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test100() As String
  Test100 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']/Edit[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']/Button[1][@Name='Open']"
End Function

Public Function Test101() As String
  Test101 = "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[2][@Name='Format Painter']"
End Function

Public Function Test102() As String
  Test102 = _
   "/Pane[1][@Name='DropShadowTop']" & vbCrLf & _
   "/Pane[2][@Name='']" & vbCrLf & _
   "/Pane[3][@Name='']" & vbCrLf & _
   "/Pane[4][@Name='']" & vbCrLf & _
   "/Pane[5][@Name='']"
End Function

Public Function Test103() As String
  Test103 = _
   "/Pane[2][@Name='']/ToolBar[1][@Name='']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']"
End Function

Public Function Test104() As String
  Test104 = "/Pane[3][@Name='']/ToolBar[1][@Name='']"
End Function

Public Function Test105() As String
  Test105 = "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']"
End Function

Public Function Test106() As String
  Test106 = "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']/MenuItem[1][@Name='System']"
End Function

Public Function Test107() As String
  Test107 = "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']/MenuItem[1][@Name='System']"
End Function

Public Function Test108() As String
  Test108 = "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']/MenuItem[1][@Name='System']"
End Function

Public Function Test109() As String
  Test109 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"
End Function

Public Function Test110() As String
  Test110 = "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[1][@Name='Cell Mode Ready']"
End Function

Public Function Test111() As String
  Test111 = _
   "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Slider[1][@Name='Zoom']/Button[1][@Name='Zoom Out']" & vbCrLf & _
   "/TitleBar[1][@Name='']/Button[1][@Name='Minimise']" & vbCrLf & _
   "/TitleBar[1][@Name='']/Button[2][@Name='Restore']" & vbCrLf & _
   "/TitleBar[1][@Name='']/Button[3][@Name='Close']"
End Function

Public Function Test112() As String
  Test112 = _
   "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[1][@Name='Cell Mode Ready']" & vbCrLf & _
   "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[2][@Name='Zoom 10%']"
End Function

Public Function Test113() As String
  Test113 = "/Edit[1][@Name='Formula Bar']"
End Function

Public Function Test114() As String
  Test114 = _
   "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[1][@Name='Cell Mode Ready']" & vbCrLf & _
   "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[2][@Name='Zoom 10%']"
End Function

Public Function Test115() As String
  Test115 = ""
End Function

Public Function Test116() As String
  Test116 = "/Edit[1][@Name='Formula Bar']"
End Function

Public Function Test117() As String
  Test117 = "/Edit[1][@Name='Formula Bar']"
End Function

Public Function Test118() As String
  Test118 = "/Edit[1][@Name='Formula Bar']"
End Function

Public Function Test119() As String
  Test119 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"
End Function

Public Function Test120() As String
  Test120 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"
End Function

Public Function Test121() As String
  Test121 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
   "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test122() As String
  Test122 = "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[2][@Name='Zoom 10%']"
End Function

Public Function Test123() As String
  Test123 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
   "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test124() As String
  Test124 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
   "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test125() As String
  Test125 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"
End Function

Public Function Test126() As String
  Test126 = "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[12][@Name='Acrobat']"
End Function

Public Function Test127() As String
  Test127 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
   "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test128() As String
  Test128 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
   "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test129() As String
  Test129 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
   "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test130() As String
  Test130 = _
   "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[1][@Name='Cell Mode Ready']" & vbCrLf & _
   "/Pane[2][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']"
End Function

Public Function Test131() As String
  Test131 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']" & vbCrLf & _
   "/Pane[4][@Name='']/ComboBox[1][@Name='Name Box']"
End Function

Public Function Test132() As String
  Test132 = _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/MenuItem[1][@Name='Customize Quick Access Toolbar']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/TitleBar[1][@Name='Excel']/MenuItem[1][@Name='Excel']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[1][@Name='Conditional Formatting']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/MenuItem[1][@Name='Format']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/MenuItem[1][@Name='More Options']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[1][@Name='Fill']" & vbCrLf & _
   "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[3][@Name='Ribbon Display Options']" & vbCrLf & _
   "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']/MenuItem[1][@Name='System']"
End Function

Public Function Test133() As String
  Test133 = "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']"
End Function

Public Function Test134() As String
  Test134 = "/Pane[3][@Name='']/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']"
End Function

Public Function TestExcel001() As String
  TestExcel001 = "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']"
End Function

Public Function TestExcel002() As String
  TestExcel002 = "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[26][@Name='A1']"
End Function

Public Function TestExcel003() As String
  TestExcel003 = _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[26][@Name='A1']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[51][@Name='A2']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[52][@Name='B2']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[54][@Name='D2']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[55][@Name='E2']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[152][@Name='B6']/@value"
End Function

Public Function TestExcel004() As String
  TestExcel004 = _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[76][@Name='A3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[77][@Name='B3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[78][@Name='C3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[79][@Name='D3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[80][@Name='E3']/@value"
End Function

Public Function TestExcel005() As String
  TestExcel005 = _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[76][@Name='A3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[77][@Name='B3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[78][@Name='C3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[79][@Name='D3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[80][@Name='E3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[101][@Name='A4']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[102][@Name='B4']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[103][@Name='C4']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[104][@Name='D4']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[105][@Name='E4']/@value"
End Function

Public Function TestExcel006() As String
  TestExcel006 = _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[126][@Name='A5']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[127][@Name='B5']/@value"
End Function

Public Function TestExcel007() As String
  TestExcel007 = _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[151][@Name='A6']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[153][@Name='C6']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[154][@Name='D6']/@value"
End Function

Public Function TestExcel008() As String
  TestExcel008 = _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[176][@Name='A7']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[177][@Name='B7']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[178][@Name='C7']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[179][@Name='D7']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[180][@Name='E7']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[181][@Name='F7']/@value"
End Function

Public Function TestExcel009() As String
  TestExcel009 = _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[76][@Name='A3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[77][@Name='B3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[78][@Name='C3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[79][@Name='D3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[80][@Name='E3']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[101][@Name='A4']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[102][@Name='B4']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[103][@Name='C4']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[104][@Name='D4']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[105][@Name='E4']/@value"
End Function




Public Function TestExcel100() As String
  TestExcel100 = "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[51][@Name='A2']"
End Function

Public Function TestExcel106() As String
   TestExcel106 = _
     "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[104][@Name='D4']/@value"
End Function

Public Function TestExcel107() As String
   TestExcel107 = _
     "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[80][@Name='E3']/@value"
End Function

Public Function TestExcel108() As String
   TestExcel108 = _
     "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[80][@Name='E3']/@value"
End Function

Public Function TestExcel109() As String
   TestExcel109 = _
     "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[54][@Name='D2']/@value"
End Function

Public Function TestExcel110() As String
   TestExcel110 = _
     "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[126][@Name='A5']/@value"
End Function

Public Function TestExcel111() As String
   TestExcel111 = _
     "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[154][@Name='D6']/@value"
End Function

Public Function TestExcel112() As String
   TestExcel112 = _
     "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[152][@Name='B6']/@value"
End Function

Public Function TestExcel113() As String
   TestExcel113 = _
     "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[55][@Name='E2']/@value"
End Function

Public Function TestExcel114() As String
  TestExcel114 = _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[54][@Name='D2']/@value"
End Function

Public Function TestExcel115() As String
  TestExcel115 = _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[54][@Name='D2']/@value"
End Function

Public Function TestExcel117() As String
  TestExcel117 = _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[51][@Name='A2']/@value" & vbCrLf & _
   "/Pane[5][@Name='']/Pane[4][@Name='" & TestWorkbookName & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[52][@Name='B2']/@value"
End Function


