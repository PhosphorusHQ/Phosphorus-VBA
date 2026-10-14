Attribute VB_Name = "HTML_RadioButtonsAndCheckboxes"
'@Folder HTML
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Function RadioButtonsAndCheckboxes() As Boolean

  Dim Succeeded As Boolean
  Succeeded = False

  Dim WebBrowser As Object
  Dim SearchBox As pLocator
  Dim SearchButton As pLocator

  On Error GoTo ErrorHandler
  
  If Not RunningAllExamples Then
    'pPath always requires the Logger class
    Phosphorus.Log4PStatic.GetLogger
'    WebBrowserCommon.GetInternetSpeeds
    WebBrowserCommon.SetDummyMobileDataInternetSpeeds
    Factory.CurrentWebBrowserType = 1
  End If
  
  Dim SourceFileURL As String
  SourceFileURL = HTMLCommon.GetULREncodedFilepath("Tests\WebPages\HTML\RadioButtonsAndCheckboxes.html")
  
  Set WebBrowser = Factory.GetNewWebBrowser
  WebBrowser.Start "HTML Radio Buttons And Checkboxes", SourceFileURL, "Radio Buttons And Checkboxes"

  Dim DocumentRoot As pLocator
  Set DocumentRoot = WebBrowser.GetRootWebArea
    
  Dim Header1 As String
  Header1 = "I Love Phosphorus ..."
  
  'Check the header
  Debug.Assert _
    DocumentRoot.FindElement("ILovePhosphorusHeader", By.NameIs, Header1) _
    .ControlType = UIAControlTypeIDs.Text

  'Check the Yes text
  With DocumentRoot.FindElement("YesRadioButtonLabel", By.NameIs, Header1, RelativeElementNumber:=2)
    Debug.Assert .ControlType = UIAControlTypeIDs.Text
    Debug.Assert .Name = "Yes"
  End With

  'Check the Yes radio button
  With DocumentRoot.FindElement("YesRadioButton", By.NameIs, Header1, RelativeElementNumber:=1)
    Debug.Assert .ControlType = UIAControlTypeIDs.RadioButton
    Debug.Assert .IsSelected
  End With

  'Check the No text
  With DocumentRoot.FindElement("NoRadioButtonLabel", By.NameIs, Header1, RelativeElementNumber:=4)
    Debug.Assert .ControlType = UIAControlTypeIDs.Text
    Debug.Assert .Name = "No"
  End With
  
  'Check the No radio button & click it
  With DocumentRoot.FindElement("NoRadioButton", By.NameIs, Header1, RelativeElementNumber:=3)
    Debug.Assert .ControlType = UIAControlTypeIDs.RadioButton
    Debug.Assert .IsNotSelected
    Debug.Assert .Click().IsSelected
  End With

  'Click the Yes radio button & check it again
  Debug.Assert _
    DocumentRoot.FindElement("YesRadioButton", By.NameIs, Header1, RelativeElementNumber:=1) _
    .Click().IsSelected
  
  Dim Header2 As String
  Header2 = "Confirm last field is disabled. Phosphorus ..."
  
  'Check the header
  Debug.Assert _
    DocumentRoot.FindElement("ConfirmLastFieldIsDisabledHeader", By.NameIs, Header2) _
    .ControlType = UIAControlTypeIDs.Text

  'Check the Rocks text
  With DocumentRoot.FindElement("RocksRadioButtonName", By.NameIs, Header2, RelativeElementNumber:=2)
    Debug.Assert .ControlType = UIAControlTypeIDs.Text
    Debug.Assert .Name = "Rocks"
  End With

  'Check the Rocks radio button
  With DocumentRoot.FindElement("RocksRadioButton", By.NameIs, Header2, RelativeElementNumber:=1)
    Debug.Assert .ControlType = UIAControlTypeIDs.RadioButton
    Debug.Assert .IsNotSelected
  End With

  'Check the Really Rocks text
  With DocumentRoot.FindElement("ReallyRocksRadioButtonName", By.NameIs, Header2, RelativeElementNumber:=4)
    Debug.Assert .ControlType = UIAControlTypeIDs.Text
    Debug.Assert .Name = "Really Rocks"
  End With

  'Check the Really Rocks radio button
  With DocumentRoot.FindElement("ReallyRocksRadioButton", By.NameIs, Header2, RelativeElementNumber:=3)
    Debug.Assert .ControlType = UIAControlTypeIDs.RadioButton
    Debug.Assert .IsNotSelected
  End With

  'Check the Maybe Rocks text
  With DocumentRoot.FindElement("MaybeRocksRadioButtonName", By.NameIs, Header2, RelativeElementNumber:=6)
    Debug.Assert .ControlType = UIAControlTypeIDs.Text
    Debug.Assert .Name = "Maybe Rocks"
  End With

  'Check the Maybe Rocks radio button
  With DocumentRoot.FindElement("MaybeRocksRadioButton", By.NameIs, Header2, RelativeElementNumber:=5)
    Debug.Assert .ControlType = UIAControlTypeIDs.RadioButton
    Debug.Assert .IsDisabled
    Debug.Assert .IsNotSelected
  End With

  'Click the Rocks radio button & check it again
  Debug.Assert _
    DocumentRoot.FindElement("RocksRadioButton", By.NameIs, Header2, RelativeElementNumber:=1) _
    .Click().IsSelected
  
  'Click the Really Rocks radio button & check it again
  Debug.Assert _
    DocumentRoot.FindElement("ReallyRocksRadioButton", By.NameIs, Header2, RelativeElementNumber:=3) _
    .Click().IsSelected

  'Check the Maybe Rocks radio button cannot be clicked
  With DocumentRoot.FindElement("MaybeRocksRadioButton", By.NameIs, Header2, RelativeElementNumber:=5)
    On Error GoTo ErrorTrap
    .Click
ErrorTrap:
    Debug.Assert Err.Number = -2147218503
    Debug.Assert Err.Description = "Element 'MaybeRocksRadioButton' is not alive!"
    On Error GoTo 0
  End With

  Dim Header3 As String
  Header3 = "Tick the check:"
  
  'Check the header
  Debug.Assert _
    DocumentRoot.FindElement("TickTheCheckHeader", By.NameIs, Header3) _
    .ControlType = UIAControlTypeIDs.Text

  'Check the Phosphorus Rocks! text
  With DocumentRoot.FindElement("TickTheCheckCheckboxName", By.NameIs, Header3, RelativeElementNumber:=2)
    Debug.Assert .ControlType = UIAControlTypeIDs.Text
    Debug.Assert .Name = "Phosphorus rocks!"
  End With

  'Check the Phosphorus Rocks! checkbox & click it
  With DocumentRoot.FindElement("TickTheCheckCheckbox", By.NameIs, Header3, RelativeElementNumber:=1)
    Debug.Assert .ControlType = UIAControlTypeIDs.CheckBox
    Debug.Assert .IsNotSelected
    Debug.Assert .Click().IsSelected
    Debug.Assert .Click().IsNotSelected
  End With

  If RunningAllExamples Then
    Debug.Print "Radio Buttons & Checkboxed Clicked!"
  Else
    MsgBox "Radio Buttons & Checkboxes Checked!"
    Phosphorus.Log4PStatic.CloseLogger
  End If
  
  Succeeded = True
  GoTo ExitSub
   
ErrorHandler:
  If RunningAllExamples Then
    Debug.Print Err.Description & " (Error Number #" & Err.Number & ") for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser in 'ExampleDomainDotCom' Example!"
  Else
    MsgBox Err.Description & " (Error Number #" & Err.Number & ") for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser in 'ExampleDomainDotCom' Example!", vbCritical
  End If
  GoTo ExitSub
  
ExitSub:
  Set WebBrowser = Nothing
  Set SearchBox = Nothing
  Set SearchButton = Nothing
  If Not RunningAllExamples Then
    Window.HighlightElements = False
  End If
  
  RadioButtonsAndCheckboxes = Succeeded

End Function
